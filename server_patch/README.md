# 中台补丁：trial 端点（试用额度）

> 归属：**海外版专属补丁**，放在海外版目录内，**不直接改动** `论文格式医生/auth-backend/`。
> 需要时由老板给 Cloudflare Token，一条命令部署。

## 为什么需要它

海外版的商业模式是 **查免费无限 + 修正第 1 次免费 + 第 2 次起付费 + 全程零水印**（模型 B）。

「第 1 次免费」必须**既能给出去、又删不掉**，否则用户删掉本地 `license.json` 就能无限白嫖。
中台现有的 `/api/activate`、`/api/heartbeat`、`/api/consume` 都是**围着激活码**转的，
没有「按机器记一次试用额度」的地方，所以补一个小端点。

## 要做什么

1. 打开 `D:\AgentSpace\论文格式医生\auth-backend\src\worker.js`
2. 在 `fetch()` 的路由区里，找到 `// ---------- 激活 ----------` 那段，**在它之前**粘贴
   `trial_endpoint.js` 里的代码块（缩进与上下文一致）
3. 部署：

```bash
cd D:/AgentSpace/论文格式医生/auth-backend
unset HTTP_PROXY HTTPS_PROXY http_proxy https_proxy
npx wrangler deploy
```

> 注意：这是**共享中台**，部署会影响国内版线上服务。请在低峰期做，并先本地
> `wrangler deploy --dry-run` 确认打包通过。

## 部署后如何自查

```bash
curl -s -X POST https://api.reedskill.com/api/trial \
  -H 'Content-Type: application/json' \
  -d '{"product":"global","machine_code":"selftest-machine","claim":false}'
# 期望：{"ok":true,"trial_used":false}
```

## 不部署会怎样

客户端 `license.py` 把 404 当作「服务器不可达」→ 走**离线宽限**（本地状态兜底）。
功能不受影响，但「删本地状态即可重置试用」这条绕过路径仍然存在。
两种状态在 `status()` 里都会如实反映（`state_ok` / `machine`），不会假报。

## 与国内版的关系

- 新增一张 `trials` 表（`product_id + machine_code` 主键），与 `codes` 表零耦合；
- 不改动任何既有路由、不改 `codes` 表结构、不影响国内版任何逻辑；
- 海外版用 `product = "global"` 的独立命名空间，与 `student` / `mentor` 互不干扰。

## 部署记录

- **已部署：2026-09-11**
  - 改动文件：`论文格式医生/auth-backend/src/worker.js`（在 `// ---------- 激活 ----------` 之前插入 35 行 trial 块，零新依赖）。
  - 部署方式：`wrangler deploy`（先 `--dry-run` 验证打包通过），Cloudflare Token 由老板现给、仅作命令内环境变量。
  - 中台提交：`54eccf6`（auth-backend 仓库，本地提交，未推送 origin）。
  - Worker Version ID：`e01c180a-39d7-41d4-9f10-044f0dc28cd2`，路由 `api.reedskill.com/*`。
  - 线上三连查全绿：查未领 `{"ok":true,"trial_used":false}` → 领取 `{"ok":true,"trial_used":true,"first":true}` → 再查 `{"ok":true,"trial_used":true}`；缺字段 `HTTP 400`；既有 `/api/heartbeat` 无回归。
  - 自测在 `trials` 表留 1 行 `global / selftest-verify-20260911`（无害，仅该虚拟机器占用一次试用）。
  - 效果：海外版「修正第 1 次免费」试用额度登记在共享中台，删本地 `license.json` / 卸载重装无法再白嫖，「试用防重置」真正闭环。
