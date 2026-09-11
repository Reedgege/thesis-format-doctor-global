// ============================================================================
// 【海外版专用补丁】trial 端点（试用额度，按机器一次性）
//
// 用途：Thesis Format Doctor Global（海外版）「修正第 1 次免费」的试用额度
//       登记在共享中台，使「删除本地 license.json / 卸载重装」无法再次白嫖。
//
// 部署方式：把本段插入 `论文格式医生/auth-backend/src/worker.js` 的
//          `fetch()` 路由区（建议紧挨 `/api/activate` 那段之前），然后
//          `npx wrangler deploy`（需老板提供 Cloudflare Token）。
//          客户端对应 `src/license/license.py` 的 `_trial_sync()`。
//
// 契约：
//   请求  POST /api/trial   { product, machine_code, claim }
//   响应  { ok:true, trial_used:true|false, first:true|false }
//         { ok:false, error:"missing_fields" }（400）
//
// 说明：
// - 表 `trials` 首次调用自动创建（幂等），无需手工建表。
// - claim=false → 只查；claim=true → 幂等领取（已领过返回 first=false，不报错）。
// - 与 codes 表无关，不影响国内版任何既有逻辑。
// ============================================================================

    // ---------- 试用额度（海外版：每台机器一次） ----------
    if (method === "POST" && path === "/api/trial") {
      const b = await request.json();
      const { product, machine_code, claim } = b;
      if (!product || !machine_code) {
        return json({ ok: false, error: "missing_fields" }, 400);
      }
      await ensureProduct(env, product);
      await env.DB.prepare(
        "CREATE TABLE IF NOT EXISTS trials (" +
        "  product_id TEXT NOT NULL," +
        "  machine_code TEXT NOT NULL," +
        "  claimed_at INTEGER," +
        "  PRIMARY KEY (product_id, machine_code)" +
        ")"
      ).run();

      const row = await env.DB.prepare(
        "SELECT * FROM trials WHERE product_id=? AND machine_code=?"
      )
        .bind(product, machine_code)
        .first();

      if (claim) {
        if (row) return json({ ok: true, trial_used: true, first: false });
        await env.DB.prepare(
          "INSERT INTO trials(product_id, machine_code, claimed_at) VALUES(?,?,?)"
        )
          .bind(product, machine_code, nowMs())
          .run();
        return json({ ok: true, trial_used: true, first: true });
      }
      return json({ ok: true, trial_used: !!row });
    }
