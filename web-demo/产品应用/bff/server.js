/**
 * BFF（Backend for Frontend）中转服务 · AI CRM 线索评分助手
 * ------------------------------------------------------------
 * 解决一句话的问题：**API Key 不该出现在浏览器里。**
 *
 * 它做三件事：
 *   1) 把 Dify 工作流接口代理到 /api/score，Key 只存在于服务器环境变量里；
 *   2) 顺手托管前端静态文件 —— 前端与接口同源，彻底没有跨域问题；
 *   3) 加两道防线：后端必填校验（前端被绕过也能兜住）+ 每 IP 限流（防止被刷爆额度）。
 *
 * 零依赖：只用 Node 内置模块（要求 Node 18+，因为用到 fetch）。
 * 启动：DIFY_API_KEY=app-xxx node server.js
 */

const http = require('http');
const fs = require('fs');
const path = require('path');
const { Readable } = require('stream');

// ---------------- 配置（全部来自环境变量，不写死在代码里） ----------------
const PORT = Number(process.env.PORT || 8787);
const DIFY_BASE = process.env.DIFY_BASE || 'https://api.dify.ai';
const DIFY_API_KEY = process.env.DIFY_API_KEY || '';
const WEB_ROOT = path.resolve(__dirname, '..');           // 静态文件根目录：产品应用/
const RATE_LIMIT = Number(process.env.RATE_LIMIT_PER_MIN || 10); // 每 IP 每分钟上限

if (!DIFY_API_KEY) {
  console.error('[BFF] 缺少环境变量 DIFY_API_KEY，无法启动。示例：DIFY_API_KEY=app-xxxx node server.js');
  process.exit(1);
}

// ---------------- 契约：必填字段（与《接口契约》保持一致，改这里也要改契约） ----------------
const REQUIRED_FIELDS = ['lead_raw', 'industry', 'company_size', 'job_title'];
const OPTIONAL_FIELDS = ['company', 'behaviors'];
const ALL_FIELDS = [...REQUIRED_FIELDS, ...OPTIONAL_FIELDS];

// ---------------- 极简限流（内存计数，重启即清空；正式环境换 Redis） ----------------
const hits = new Map();
function rateLimited(ip) {
  const now = Date.now();
  const rec = hits.get(ip) || { count: 0, ts: now };
  if (now - rec.ts > 60_000) { rec.count = 0; rec.ts = now; }
  rec.count += 1;
  hits.set(ip, rec);
  return rec.count > RATE_LIMIT;
}

// ---------------- 错误翻译：把 Dify 的报错翻成人话 ----------------
function humanize(status, text) {
  if (status === 401) return { code: 'AUTH_FAILED', message: '密钥失效或未授权，请联系维护者' };
  if (text && text.includes('in input form')) {
    const missing = (text.match(/([\w]+) is required/) || [])[1] || '未知字段';
    return { code: 'MISSING_FIELD', message: `缺少必填字段：${missing}` };
  }
  if (text && text.includes('quota')) {
    return { code: 'QUOTA_EXCEEDED', message: '模型额度已用完，请稍后再试或切换模型供应商' };
  }
  if (text && text.includes('conversation')) {
    return { code: 'UPSTREAM_ERROR', message: '上游应用类型不匹配（本接口只代理 workflow 应用）' };
  }
  return { code: 'UPSTREAM_ERROR', message: `上游返回异常（HTTP ${status}）` };
}

// ---------------- 静态文件服务（只读、防目录穿越） ----------------
const MIME = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8', '.json': 'application/json; charset=utf-8',
  '.svg': 'image/svg+xml', '.png': 'image/png', '.jpg': 'image/jpeg', '.ico': 'image/x-icon' };

function serveStatic(req, res) {
  const urlPath = decodeURIComponent(new URL(req.url, 'http://x').pathname);
  const rel = urlPath === '/' ? 'index.html' : urlPath.replace(/^\/+/, '');
  const filePath = path.resolve(WEB_ROOT, rel);
  if (!filePath.startsWith(WEB_ROOT)) { res.writeHead(403).end('Forbidden'); return; }
  fs.readFile(filePath, (err, buf) => {
    if (err) { res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' }).end('404 Not Found'); return; }
    res.writeHead(200, { 'Content-Type': MIME[path.extname(filePath).toLowerCase()] || 'application/octet-stream' });
    res.end(buf);
  });
}

// ---------------- 核心：代理 /api/score → Dify /v1/workflows/run ----------------
async function proxyScore(req, res) {
  const ip = req.socket.remoteAddress || 'unknown';
  if (rateLimited(ip)) {
    res.writeHead(429, { 'Content-Type': 'application/json; charset=utf-8' })
       .end(JSON.stringify({ code: 'RATE_LIMITED', message: `调用过于频繁（上限 ${RATE_LIMIT} 次/分钟），请稍后再试` }));
    return;
  }

  // 读取请求体
  const chunks = [];
  for await (const c of req) chunks.push(c);
  let payload;
  try { payload = JSON.parse(Buffer.concat(chunks).toString('utf8') || '{}'); }
  catch { res.writeHead(400, { 'Content-Type': 'application/json; charset=utf-8' })
             .end(JSON.stringify({ code: 'BAD_JSON', message: '请求体不是合法 JSON' })); return; }

  // 后端兜底校验：前端校验能被绕过，这里必须再拦一次
  const inputs = payload.inputs || {};
  const missing = REQUIRED_FIELDS.filter((f) => !String(inputs[f] || '').trim());
  if (missing.length) {
    res.writeHead(400, { 'Content-Type': 'application/json; charset=utf-8' })
       .end(JSON.stringify({ code: 'MISSING_FIELD', message: `缺少必填字段：${missing.join('、')}` }));
    return;
  }
  // 只透传契约字段，防止前端塞脏参数
  const clean = {};
  for (const f of ALL_FIELDS) if (inputs[f] !== undefined) clean[f] = String(inputs[f]);

  const t0 = Date.now();
  let upstream;
  try {
    upstream = await fetch(`${DIFY_BASE}/v1/workflows/run`, {
      method: 'POST',
      headers: { 'Authorization': `Bearer ${DIFY_API_KEY}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({
        inputs: clean,
        response_mode: 'streaming',                 // 长流程必须流式，blocking 有 100s 上限
        user: payload.user || `bff-${ip}`,
      }),
    });
  } catch (e) {
    res.writeHead(502, { 'Content-Type': 'application/json; charset=utf-8' })
       .end(JSON.stringify({ code: 'NETWORK_ERROR', message: `无法连接上游：${e.message}` }));
    return;
  }

  if (!upstream.ok) {
    const text = await upstream.text().catch(() => '');
    const err = humanize(upstream.status, text);
    console.error(`[BFF] 上游 ${upstream.status} ${err.code}: ${text.slice(0, 200)}`);
    res.writeHead(upstream.status, { 'Content-Type': 'application/json; charset=utf-8' })
       .end(JSON.stringify(err));
    return;
  }

  // 透传 SSE（打字机效果 + 节点级进度），并在结束时打印一条访问日志
  res.writeHead(200, { 'Content-Type': 'text/event-stream; charset=utf-8', 'Cache-Control': 'no-cache', 'Connection': 'keep-alive' });
  const stream = Readable.fromWeb(upstream.body);
  stream.pipe(res);
  stream.on('end', () => console.log(`[BFF] ${ip} 评分完成，耗时 ${((Date.now() - t0) / 1000).toFixed(1)}s`));
  stream.on('error', () => res.end());
}

// ---------------- 路由 ----------------
const server = http.createServer((req, res) => {
  if (req.method === 'POST' && req.url.split('?')[0] === '/api/score') return proxyScore(req, res);
  if (req.method === 'GET' && req.url.split('?')[0] === '/healthz') {
    res.writeHead(200, { 'Content-Type': 'application/json' }).end(JSON.stringify({ ok: true }));
    return;
  }
  if (req.method === 'GET') return serveStatic(req, res);
  res.writeHead(405).end('Method Not Allowed');
});

server.listen(PORT, () => {
  console.log(`[BFF] 已启动：http://localhost:${PORT}`);
  console.log(`[BFF] 前端静态目录：${WEB_ROOT}`);
  console.log(`[BFF] 接口：POST /api/score ｜ 限流：${RATE_LIMIT} 次/分钟/IP ｜ Key：已注入（不下发浏览器）`);
});
