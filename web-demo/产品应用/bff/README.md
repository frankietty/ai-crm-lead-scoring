# BFF 后端中转 · 最小可运行示例

> 解决 Demo 的唯一「不可交付项」：**API Key 暴露在前端源码里。**
> 本目录是「正式交付形态」的参考实现——把密钥搬到服务器，前端只跟自己服务器说话。

---

## 一、它和 Demo 形态的差别

| | Demo 形态（当前前端） | BFF 形态（本目录） |
|---|---|---|
| Key 位置 | 浏览器 `config.js`，**F12 就能看到** | 服务器环境变量，**从不下发** |
| 跨域 | 直连 `api.dify.ai`，受 CORS 限制 | 同源，**无跨域** |
| 防刷 | 无，谁都能刷爆你的模型额度 | 每 IP 每分钟限流（默认 10 次） |
| 参数校验 | 只在前端，可绕过 | 服务器兜底拦一次 + 字段白名单 |
| 可观测 | 只能看 Dify 后台 | 服务器日志记耗时、错误码、来源 IP |

---

## 二、本地跑起来（2 分钟）

```bash
cd web-demo/产品应用/bff
DIFY_API_KEY=app-你的密钥 node server.js
# 打开 http://localhost:8787  → 就是三页产品应用（index / workbench / guide）
# 健康检查：http://localhost:8787/healthz
```

启动后：
- 静态页面从 `../`（产品应用目录）托管，所以**访问 8787 就能打开 index.html**；
- 评分接口变成同源 `POST /api/score`。

> Windows PowerShell 设环境变量：`$env:DIFY_API_KEY="app-xxx"; node server.js`

---

## 三、前端要改哪一行

把页面里直连 Dify 的地址换成同源接口即可：

```js
// 改前（Demo 形态）：直连 Dify，Key 在前端
const endpoint = 'https://api.dify.ai/v1/workflows/run';
headers['Authorization'] = `Bearer ${window.DIFY_API_KEY}`;

// 改后（BFF 形态）：只跟自己的服务器说话
const endpoint = '/api/score';
// 不再需要 Authorization 头；请求体仍是 { inputs: {...}, user: 'xxx' }
```

返回格式不变（SSE 流式事件原样透传），**所以页面其余逻辑一行都不用改**。

---

## 四、部署到云（三条路，任选）

| 方式 | 做法 | 备注 |
|---|---|---|
| 云函数 / Serverless | 把 `server.js` 的 handler 改成函数入口，环境变量配 `DIFY_API_KEY` | 最省事，按量计费 |
| 容器 | 加一个 `FROM node:22-alpine` 的 Dockerfile，`CMD ["node","server.js"]` | 便于长期运行 |
| 自己的服务器 | `pm2 start server.js --name crm-bff` | 记得配 Nginx 反代 + HTTPS |

**上线后必须做的三件事**：① 把 `DIFY_API_KEY` 配成平台密钥而非明文；② 前置 HTTPS；③ 把内存限流换成 Redis（多实例共享计数）。

---

## 五、安全清单（交付前逐条确认）

- [ ] `DIFY_API_KEY` 只存在于服务器环境变量，**仓库里搜不到**
- [ ] 前端全部改走 `/api/score`，源码里不再出现 `api.dify.ai`
- [ ] 已开启限流，并验证被限流时返回的是**人话**（429 + 中文提示），不是白屏
- [ ] 后端必填校验生效：直接 `curl` 发一个空 `inputs`，应返回 `MISSING_FIELD`
- [ ] 日志不打印 Key、不打印客户完整留言（只记耗时与状态）
- [ ] 已加 HTTPS，且前端与接口同域（无混合内容告警）

---

## 六、为什么这一步值得做（面试/客户可直接说）

> 「Demo 阶段我把 Key 放前端，是为了让您点开链接就能用；正式交付我会加一层 BFF——密钥留在服务器，前端只调自己的接口，再叠上限流、日志和鉴权。**这不是我会不会写的问题，是我知道 POC 和生产之间差哪一道门。**」
