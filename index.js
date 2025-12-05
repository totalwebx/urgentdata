// server.js
require('dotenv').config();
const express = require('express');
const sql = require('mssql');
const cors = require('cors');
const http = require('http');                 // ⬅️ added
const { Server } = require('socket.io');      // ⬅️ added

const app = express();
app.use(cors());
app.use(express.json());

const PORT = process.env.PORT || 3000;
const HOST = process.env.HOST || "0.0.0.0";


/* ===================== Socket.IO ===================== */
const server = http.createServer(app);        // ⬅️ wrap Express in HTTP server
const io = new Server(server, {
  cors: { origin: '*', methods: ['GET','POST','PATCH','PUT','DELETE'] }
});
app.set('io', io);                            // ⬅️ so we can grab io in handlers

io.on('connection', (socket) => {
  // Optional: targeted rooms by machine/type if you need later
  // socket.on('subscribe', ({ machine, type }) => {
  //   if (machine) socket.join(`mc:${String(machine).toUpperCase()}`);
  //   if (type) socket.join(`type:${String(type).toLowerCase()}`);
  // });
});

/* ===================== DB config ===================== */
const config = {
  user: process.env.DB_USER || 'Test',
  password: process.env.DB_PASS || 'Welcome@2021',
  server: process.env.DB_HOST || '10.71.5.26',
  database: process.env.DB_NAME || 'Précontrole',
  options: { encrypt: false, trustServerCertificate: true },
  pool: { max: 10, min: 1, idleTimeoutMillis: 30000 }
};

// ---- single, shared pool ----
let poolPromise;
function getPool() {
  if (!poolPromise) poolPromise = new sql.ConnectionPool(config).connect();
  return poolPromise;
}

// ---- helpers ----
function toInt(v, def) {
  const n = parseInt(v, 10);
  return Number.isFinite(n) && n > 0 ? n : def;
}
function ok(value) {
  return value !== undefined && value !== null && String(value).trim() !== '';
}

/* -------------------------------------------------
 Filters:
  ?status=OK|NOK
  ?machine=MC27
  ?machines=MC1,MC2,MC3
  ?machineLike=MC
  ?unico=LH-148031
  ?from=2025-10-21T00:00:00
  ?to=2025-10-22T23:59:59
  ?declaredBy=588|AISSAM
  ?correctedBy=1935|HANIFA
------------------------------------------------- */
app.get('/urgent', async (req, res) => {
  try {
    const pool = await getPool();
    const conditions = [];
    const request = pool.request();

    // helpers
    const toValidDate = (s) => {
      const d = new Date(s);
      return isNaN(d.getTime()) ? null : d;
    };

    // ---- simple filters
    if (ok(req.query.status)) {
      request.input('status', sql.VarChar, String(req.query.status).trim().toUpperCase());
      conditions.push('UPPER(urg.[Statut]) = @status');
    }

    if (ok(req.query.unico)) {
      request.input('unico', sql.VarChar, `%${req.query.unico.trim()}%`);
      conditions.push('urg.[Unico] LIKE @unico');
    }

    if (ok(req.query.from)) {
      const d = toValidDate(req.query.from);
      if (d) {
        request.input('from', sql.DateTime2, d);
        conditions.push('urg.[Date_Declaration] >= @from');
      }
    }

    if (ok(req.query.to)) {
      const d = toValidDate(req.query.to);
      if (d) {
        request.input('to', sql.DateTime2, d);
        conditions.push('urg.[Date_Declaration] <= @to');
      }
    }

    // 🔍 who declared
    if (ok(req.query.declaredBy)) {
      const val = `%${req.query.declaredBy.trim()}%`;
      request.input('declaredBy', sql.VarChar, val);
      conditions.push(`
        (
          LTRIM(RTRIM(urg.[Declarer_Par])) LIKE @declaredBy OR
          uDecl.[Nom]    LIKE @declaredBy OR
          uDecl.[Prenom] LIKE @declaredBy
        )
      `);
    }

    // 🔍 who corrected
    if (ok(req.query.correctedBy)) {
      const val = `%${req.query.correctedBy.trim()}%`;
      request.input('correctedBy', sql.VarChar, val);
      conditions.push(`
        (
          LTRIM(RTRIM(urg.[Corriger_Par])) LIKE @correctedBy OR
          uCorr.[Nom]    LIKE @correctedBy OR
          uCorr.[Prenom] LIKE @correctedBy
        )
      `);
    }

    // ---- robust machine filters
    if (ok(req.query.machines)) {
      const list = String(req.query.machines)
        .split(',')
        .map(s => s.trim())
        .filter(Boolean);

      if (list.length) {
        const orParts = [];
        list.forEach((m, i) => {
          const key = `mach${i}`;
          request.input(key, sql.VarChar, m.toUpperCase().trim());
          orParts.push(`UPPER(LTRIM(RTRIM(REPLACE(REPLACE(urg.[Machine], CHAR(160), ' '), CHAR(9), ' ')))) = @${key}`);
        });
        conditions.push(`(${orParts.join(' OR ')})`);
      }
    } else if (ok(req.query.machine)) {
      request.input('machine', sql.VarChar, req.query.machine.toUpperCase().trim());
      conditions.push(`UPPER(LTRIM(RTRIM(REPLACE(REPLACE(urg.[Machine], CHAR(160), ' '), CHAR(9), ' ')))) = @machine`);
    }

    if (ok(req.query.machineLike)) {
      request.input('machineLike', sql.VarChar, `%${req.query.machineLike.toUpperCase().trim()}%`);
      conditions.push(`UPPER(LTRIM(RTRIM(REPLACE(REPLACE(urg.[Machine], CHAR(160), ' '), CHAR(9), ' ')))) LIKE @machineLike`);
    }

    const whereSql = conditions.length ? `WHERE ${conditions.join(' AND ')}` : '';

    // JOIN users (declarer + corrector)
    const query = `
      SELECT
        urg.*,
        uDecl.[Nom]     AS Decl_Nom,
        uDecl.[Prenom]  AS Decl_Prenom,
        uDecl.[Role]    AS Decl_Role,
        uDecl.[Mlle]    AS Decl_Badge,
        uCorr.[Nom]     AS Corr_Nom,
        uCorr.[Prenom]  AS Corr_Prenom,
        uCorr.[Role]    AS Corr_Role,
        uCorr.[Mlle]    AS Corr_Badge
      FROM [dbo].[M5_Urgent] AS urg
      LEFT JOIN [dbo].[M5_Users] AS uDecl
        ON LTRIM(RTRIM(uDecl.[Mlle])) = LTRIM(RTRIM(urg.[Declarer_Par]))
      LEFT JOIN [dbo].[M5_Users] AS uCorr
        ON LTRIM(RTRIM(uCorr.[Mlle])) = LTRIM(RTRIM(urg.[Corriger_Par]))
      ${whereSql}
      ORDER BY urg.[Date_Declaration] DESC;
    `;

    const rows = await request.query(query);

    const results = rows.recordset.map(r => {
      const base = {
        id: r.id,
        unico: r.Unico?.trim(),
        declaredAt: r.Date_Declaration,
        correctedAt: r.Date_Correction,
        status: r.Statut,
        machine: r.Machine?.trim() || null,
        planB: r.Plan_B?.trim() || null,
        mcPb: r.McPb ?? null,
        type: r.Type ? String(r.Type).toLowerCase() : null,
        timeRemaining: r.Temps_Restant?.trim() || null,
      };

      base.declaredBy = {
        matricule: r.Declarer_Par ? String(r.Declarer_Par).trim() : null,
        "full name": (r.Decl_Prenom || r.Decl_Nom)
          ? `${(r.Decl_Prenom || '').trim()} ${(r.Decl_Nom || '').trim()}`.trim()
          : null,
        role: r.Decl_Role ? String(r.Decl_Role).trim() : null
      };

      if (r.Corriger_Par) {
        base.correctedBy = {
          matricule: String(r.Corriger_Par).trim(),
          "full name": (r.Corr_Prenom || r.Corr_Nom)
            ? `${(r.Corr_Prenom || '').trim()} ${(r.Corr_Nom || '').trim()}`.trim()
            : null,
          role: r.Corr_Role ? String(r.Corr_Role).trim() : null
        };
      } else {
        base.correctedBy = null;
      }

      return base;
    });

    res.json({ count: results.length, results });
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: 'Failed to fetch urgent data' });
  }
});

// Distinct machines from M5_Urgent (normalized, filtered like /urgent, but no machine filters)
app.get('/urgent/machines', async (req, res) => {
  try {
    const pool = await getPool();
    const conditions = [];
    const request = pool.request();

    const ok = v => v !== undefined && v !== null && String(v).trim() !== '';
    const toValidDate = (s) => {
      const d = new Date(s);
      return isNaN(d.getTime()) ? null : d;
    };

    // copy the common filters (NO machine filters here!)
    if (ok(req.query.status)) {
      request.input('status', sql.VarChar, String(req.query.status).trim().toUpperCase());
      conditions.push('UPPER(urg.[Statut]) = @status');
    }
    if (ok(req.query.unico)) {
      request.input('unico', sql.VarChar, `%${req.query.unico.trim()}%`);
      conditions.push('urg.[Unico] LIKE @unico');
    }
    if (ok(req.query.from)) {
      const d = toValidDate(req.query.from);
      if (d) {
        request.input('from', sql.DateTime2, d);
        conditions.push('urg.[Date_Declaration] >= @from');
      }
    }
    if (ok(req.query.to)) {
      const d = toValidDate(req.query.to);
      if (d) {
        request.input('to', sql.DateTime2, d);
        conditions.push('urg.[Date_Declaration] <= @to');
      }
    }
    if (ok(req.query.declaredBy)) {
      const v = `%${req.query.declaredBy.trim()}%`;
      request.input('declaredBy', sql.VarChar, v);
      conditions.push(`
        (
          LTRIM(RTRIM(urg.[Declarer_Par])) LIKE @declaredBy OR
          uDecl.[Nom]    LIKE @declaredBy OR
          uDecl.[Prenom] LIKE @declaredBy
        )
      `);
    }
    if (ok(req.query.correctedBy)) {
      const v = `%${req.query.correctedBy.trim()}%`;
      request.input('correctedBy', sql.VarChar, v);
      conditions.push(`
        (
          LTRIM(RTRIM(urg.[Corriger_Par])) LIKE @correctedBy OR
          uCorr.[Nom]    LIKE @correctedBy OR
          uCorr.[Prenom] LIKE @correctedBy
        )
      `);
    }

    const whereSql = conditions.length ? `WHERE ${conditions.join(' AND ')}` : '';

    // Normalize Machine
    const q = await request.query(`
      ;WITH F AS (
        SELECT urg.[Machine]
        FROM [dbo].[M5_Urgent] AS urg
        LEFT JOIN [dbo].[M5_Users] AS uDecl
          ON LTRIM(RTRIM(uDecl.[Mlle])) = LTRIM(RTRIM(urg.[Declarer_Par]))
        LEFT JOIN [dbo].[M5_Users] AS uCorr
          ON LTRIM(RTRIM(uCorr.[Mlle])) = LTRIM(RTRIM(urg.[Corriger_Par]))
        ${whereSql}
      ),
      N AS (
        SELECT
          UPPER(LTRIM(RTRIM(REPLACE(REPLACE([Machine], CHAR(160), ' '), CHAR(9), ' ')))) AS MNorm
        FROM F
        WHERE [Machine] IS NOT NULL
          AND LTRIM(RTRIM([Machine])) <> ''
          AND [Machine] <> '-'
      )
      SELECT DISTINCT MNorm AS Machine
      FROM N
      WHERE MNorm IS NOT NULL AND MNorm <> ''
      ORDER BY Machine ASC;
    `);

    const machines = (q.recordset || []).map(r => r.Machine);
    res.json({ count: machines.length, results: machines });
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: 'Failed to fetch machines' });
  }
});

/* ---------------------------------------------------------------
POST /urgent  (emit urgent:added)
--------------------------------------------------------------- */
app.post('/urgent', async (req, res) => {
  try {
    const pool = await getPool();
    let { declarerMatricule, password, urgents, unico, machine, planB, mcPb, type, timeRemaining } = req.body || {};

    if (!Array.isArray(urgents)) {
      if (!unico || !machine)
        return res.status(400).json({ error: 'Missing unico or machine.' });
      urgents = [{ unico, machine, planB, mcPb, type, timeRemaining }];
    }

    if (!declarerMatricule || !password)
      return res.status(400).json({ error: 'Required: declarerMatricule and password.' });

    const auth = pool.request();
    auth.input('mlle', sql.VarChar, String(declarerMatricule).trim());
    auth.input('pass', sql.VarChar, String(password).trim());
    const userRes = await auth.query(`
      SELECT [Mlle],[Nom],[Prenom],[Role]
      FROM [dbo].[M5_Users]
      WHERE REPLACE(LTRIM(RTRIM([Mlle])),' ','') = REPLACE(@mlle,' ','')
        AND LTRIM(RTRIM([Password])) = LTRIM(RTRIM(@pass))
    `);
    if (!userRes.recordset.length)
      return res.status(401).json({ error: 'Invalid matricule or password.' });

    const user = userRes.recordset[0];

    const allUnicos = urgents.map(u => String(u.unico || '').trim()).filter(Boolean);
    if (!allUnicos.length)
      return res.status(400).json({ error: 'No valid Unico provided.' });

    const existing = await pool.request().query(`
      SELECT DISTINCT [Unico]
      FROM [dbo].[M5_Wires]
      WHERE [Unico] IN (${allUnicos.map(u => `'${u.replace(/'/g, "''")}'`).join(',')});
    `);

    const foundUnicos = new Set(existing.recordset.map(r => r.Unico?.trim()));
    const missing = allUnicos.filter(u => !foundUnicos.has(u));
    if (missing.length > 0) {
      return res.status(400).json({ error: 'Some Unico values do not exist in wires.', missing });
    }

    const results = [];
    for (const u of urgents) {
      if (!u.unico || !u.machine) continue;
      const declAt = new Date();

      const ins = pool.request();
      ins.input('Unico', sql.NVarChar, String(u.unico).trim());
      ins.input('Machine', sql.NVarChar, String(u.machine).trim());
      ins.input('Declarer_Par', sql.VarChar, String(user.Mlle).trim());
      ins.input('Statut', sql.VarChar, 'NOK');
      ins.input('Date_Declaration', sql.DateTime2, declAt);
      ins.input('Plan_B', sql.NVarChar, u.planB ? String(u.planB).trim() : null);
      ins.input('McPb', sql.NVarChar, u.mcPb ? String(u.mcPb).trim() : null);
      ins.input('Type', sql.NVarChar, u.type ? String(u.type).trim() : null);
      ins.input('Temps_Restant', sql.NVarChar, u.timeRemaining ? String(u.timeRemaining).trim() : null);

      const inserted = await ins.query(`
        INSERT INTO [dbo].[M5_Urgent]
          ([Unico],[Date_Declaration],[Declarer_Par],[Statut],[Machine],
           [Plan_B],[McPb],[Type],[Temps_Restant])
        OUTPUT INSERTED.id
        VALUES (@Unico,@Date_Declaration,@Declarer_Par,@Statut,@Machine,
                @Plan_B,@McPb,@Type,@Temps_Restant);
      `);

      const id = inserted.recordset?.[0]?.id;
      if (!id) continue;

      const row = {
        id,
        unico: u.unico.trim(),
        declaredAt: declAt,
        status: 'NOK',
        machine: u.machine.trim(),
        planB: u.planB ?? null,
        mcPb: u.mcPb ?? null,
        type: u.type ? u.type.toLowerCase().trim() : null,
        timeRemaining: u.timeRemaining ?? null,
        declaredBy: {
          matricule: user.Mlle.trim(),
          "full name": `${(user.Prenom || '').trim()} ${(user.Nom || '').trim()}`.trim(),
          role: (user.Role || '').trim()
        },
        correctedBy: null,
        correctedAt: null
      };
      results.push(row);

      // 🔔 notify all clients
      io.emit('urgent:added', {
        id: row.id,
        unico: row.unico,
        machine: row.machine,
        type: row.type,
        declaredAt: row.declaredAt,
        by: row.declaredBy?.matricule
      });
      // Optional: targeted rooms
      // io.to(`mc:${row.machine.toUpperCase()}`).emit('urgent:added', row);
      // if (row.type) io.to(`type:${row.type}`).emit('urgent:added', row);
    }

    if (!results.length)
      return res.status(400).json({ error: 'No valid urgents inserted.' });

    res.status(201).json({ count: results.length, results });
  } catch (err) {
    console.error('❌ POST /urgent error:', err);
    res.status(500).json({ error: 'Failed to declare urgent(s).' });
  }
});

/* ---------------------------------------------------------------
PATCH /urgent/planb  (emit urgent:planb)
---------------------------------------------------------------- */
app.patch('/urgent/planb', async (req, res) => {
  try {
    const pool = await getPool();
    const { unico, McPb } = req.body || {};

    if (!unico || !McPb)
      return res.status(400).json({ error: 'Required: unico and McPb.' });

    const find = pool.request();
    find.input('unico', sql.NVarChar, unico.trim());
    const found = await find.query(`
      SELECT TOP 1 id, Machine, Type FROM [dbo].[M5_Urgent]
      WHERE LTRIM(RTRIM([Unico])) = LTRIM(RTRIM(@unico))
        AND UPPER([Statut]) = 'NOK'
      ORDER BY [Date_Declaration] DESC, [id] DESC;
    `);

    if (!found.recordset.length)
      return res.status(404).json({ error: 'No NOK urgent found for this Unico.' });

    const row = found.recordset[0];

    const upd = pool.request();
    upd.input('id', sql.Int, row.id);
    upd.input('mcPb', sql.NVarChar, McPb.trim());
    await upd.query(`
      UPDATE [dbo].[M5_Urgent]
      SET [Plan_B] = 1, [McPb] = @mcPb
      WHERE [id] = @id;
    `);

    // 🔔 notify all clients
    io.emit('urgent:planb', {
      id: row.id,
      unico,
      machine: row.Machine?.trim() || null,
      type: row.Type ? String(row.Type).toLowerCase() : null,
      mcPb: McPb.trim()
    });

    return res.json({ success: true, unico, McPb, Plan_B: true });
  } catch (err) {
    console.error('❌ PATCH /urgent/planb error:', err);
    res.status(500).json({ error: 'Failed to set Plan B.' });
  }
});

/* ---------------------------------------------------------------
PATCH /urgent/resolve  (emit urgent:resolved)
--------------------------------------------------------------- */
app.patch('/urgent/resolve', async (req, res) => {
  try {
    const pool = await getPool();
    const { unico, unicos, correctorMatricule, password } = req.body || {};

    if (!correctorMatricule || !password) {
      return res.status(400).json({ error: 'Required: correctorMatricule, password.' });
    }

    let list = Array.isArray(unicos) ? unicos : (unico ? [unico] : []);
    list = list.map(u => String(u || '').trim()).filter(Boolean);
    if (!list.length) {
      return res.status(400).json({ error: 'Provide "unico" or "unicos" (non-empty).' });
    }

    const auth = pool.request();
    auth.input('mlle', sql.VarChar, String(correctorMatricule).trim());
    auth.input('pass', sql.VarChar, String(password).trim());
    const authRes = await auth.query(`
      SELECT [Mlle],[Nom],[Prenom],[Role]
      FROM [dbo].[M5_Users]
      WHERE REPLACE(LTRIM(RTRIM([Mlle])),' ','') = REPLACE(@mlle,' ','')
        AND LTRIM(RTRIM([Password])) = LTRIM(RTRIM(@pass))
    `);
    if (!authRes.recordset.length) {
      return res.status(401).json({ error: 'Invalid corrector credentials.' });
    }
    const corrector = authRes.recordset[0];

    const results = [];
    const skipped = [];

    for (const u of list) {
      const find = pool.request();
      find.input('unico', sql.NVarChar, u);
      const found = await find.query(`
        SELECT TOP 1 * FROM [dbo].[M5_Urgent]
        WHERE LTRIM(RTRIM([Unico])) = LTRIM(RTRIM(@unico)) AND UPPER([Statut]) = 'NOK'
        ORDER BY [Date_Declaration] DESC, [id] DESC;
      `);
      if (!found.recordset.length) {
        skipped.push({ unico: u, reason: 'No NOK urgent found' });
        continue;
      }

      const target = found.recordset[0];
      const now = new Date();

      const upd = pool.request();
      upd.input('id', sql.Int, target.id);
      upd.input('corrBy', sql.VarChar, String(corrector.Mlle).trim());
      upd.input('corrAt', sql.DateTime2, now);
      await upd.query(`
        UPDATE [dbo].[M5_Urgent]
        SET [Statut] = 'OK',
            [Corriger_Par] = @corrBy,
            [Date_Correction] = @corrAt
        WHERE [id] = @id;
      `);

      const sel = pool.request();
      sel.input('id', sql.Int, target.id);
      const rset = await sel.query(`
        SELECT
          urg.*,
          uDecl.[Nom]   AS Decl_Nom,  uDecl.[Prenom] AS Decl_Prenom,  uDecl.[Role] AS Decl_Role,
          uCorr.[Nom]   AS Corr_Nom,  uCorr.[Prenom] AS Corr_Prenom,  uCorr.[Role] AS Corr_Role
        FROM [dbo].[M5_Urgent] AS urg
        LEFT JOIN [dbo].[M5_Users] AS uDecl
          ON LTRIM(RTRIM(uDecl.[Mlle])) = LTRIM(RTRIM(urg.[Declarer_Par]))
        LEFT JOIN [dbo].[M5_Users] AS uCorr
          ON LTRIM(RTRIM(uCorr.[Mlle])) = LTRIM(RTRIM(urg.[Corriger_Par]))
        WHERE urg.[id] = @id;
      `);

      const r = rset.recordset[0];
      const row = {
        id: r.id,
        unico: r.Unico?.trim(),
        declaredAt: r.Date_Declaration,
        correctedAt: r.Date_Correction,
        status: r.Statut,
        machine: r.Machine?.trim() || null,
        planB: r.Plan_B?.trim() || null,
        mcPb: r.McPb ?? null,
        type: r.Type ? String(r.Type).toLowerCase() : null,
        timeRemaining: r.Temps_Restant?.trim() || null,
        declaredBy: {
          matricule: r.Declarer_Par ? String(r.Declarer_Par).trim() : null,
          "full name": (r.Decl_Prenom || r.Decl_Nom)
            ? `${(r.Decl_Prenom || '').trim()} ${(r.Decl_Nom || '').trim()}`.trim()
            : null,
          role: r.Decl_Role ? String(r.Decl_Role).trim() : null
        },
        correctedBy: {
          matricule: String(corrector.Mlle).trim(),
          "full name": `${(corrector.Prenom || '').trim()} ${(corrector.Nom || '').trim()}`.trim(),
          role: (corrector.Role || '').trim()
        }
      };
      results.push(row);

      // 🔔 notify all clients
      io.emit('urgent:resolved', {
        id: row.id,
        unico: row.unico,
        machine: row.machine,
        type: row.type,
        correctedAt: row.correctedAt,
        by: row.correctedBy?.matricule
      });
      // Optional rooms:
      // io.to(`mc:${row.machine?.toUpperCase()}`).emit('urgent:resolved', row);
      // if (row.type) io.to(`type:${row.type}`).emit('urgent:resolved', row);
    }

    return res.json({ count: results.length, results, skipped });
  } catch (e) {
    console.error('❌ PATCH /urgent/resolve (batch) error:', e);
    res.status(500).json({ error: 'Failed to resolve urgent(s).' });
  }
});

/* ---------------------------------------------------------------
GET /users (with counts)
--------------------------------------------------------------- */
app.get('/users', async (req, res) => {
  try {
    const pool = await getPool();
    const conditions = [];
    const request = pool.request();

    if (ok(req.query.role)) {
      request.input('role', sql.VarChar, req.query.role.trim());
      conditions.push('[Role] = @role');
    }
    if (ok(req.query.q)) {
      const like = `%${req.query.q.trim()}%`;
      request.input('q1', sql.VarChar, like);
      request.input('q2', sql.VarChar, like);
      request.input('q3', sql.VarChar, like);
      conditions.push('([Nom] LIKE @q1 OR [Prenom] LIKE @q2 OR [Mlle] LIKE @q3)');
    }

    const whereSql = conditions.length ? `WHERE ${conditions.join(' AND ')}` : '';

    const query = `
      SELECT 
        u.[Nom],
        u.[Prenom],
        u.[Mlle],
        u.[Password],
        u.[Role],
        (SELECT COUNT(*) FROM [dbo].[M5_Urgent] WHERE LTRIM(RTRIM([Declarer_Par])) = LTRIM(RTRIM(u.[Mlle]))) AS declaredCount,
        (SELECT COUNT(*) FROM [dbo].[M5_Urgent] WHERE LTRIM(RTRIM([Corriger_Par])) = LTRIM(RTRIM(u.[Mlle]))) AS correctedCount
      FROM [dbo].[M5_Users] AS u
      ${whereSql}
      ORDER BY u.[Nom] ASC, u.[Prenom] ASC;
    `;

    const rows = await request.query(query);

    const users = rows.recordset.map(r => ({
      matricule: r.Mlle?.trim(),
      fullName: `${(r.Prenom || '').trim()} ${(r.Nom || '').trim()}`.trim(),
      role: r.Role?.trim(),
      password: r.Password?.trim() || null,
      declaredCount: r.declaredCount ?? 0,
      correctedCount: r.correctedCount ?? 0
    }));

    res.json({ count: users.length, results: users });
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: 'Failed to fetch users' });
  }
});

/* ---------------------------------------------------------------
POST /users
--------------------------------------------------------------- */
app.post('/users', async (req, res) => {
  try {
    const pool = await getPool();
    const { matricule, firstName, lastName, role, password } = req.body || {};

    if (!matricule || !firstName || !lastName || !role || !password) {
      return res.status(400).json({
        error: 'Required: matricule, firstName, lastName, role, password.'
      });
    }

    const roleWhitelist = ['Admin', 'Cutting', 'Opera', 'Alimentation'];
    if (!roleWhitelist.includes(String(role).trim())) {
      return res.status(400).json({ error: `Invalid role. Allowed: ${roleWhitelist.join(', ')}` });
    }

    const mlleCanon = String(matricule).trim();
    const chk = pool.request();
    chk.input('mlle', sql.VarChar, mlleCanon);
    const exists = await chk.query(`
      SELECT 1 AS x
      FROM [dbo].[M5_Users]
      WHERE REPLACE(LTRIM(RTRIM([Mlle])),' ','') = REPLACE(@mlle,' ','')
    `);
    if (exists.recordset.length) {
      return res.status(409).json({ error: 'Matricule already exists.' });
    }

    const ins = pool.request();
    ins.input('nom', sql.NVarChar, String(lastName).trim());
    ins.input('prenom', sql.NVarChar, String(firstName).trim());
    ins.input('mlle', sql.VarChar, mlleCanon);
    ins.input('pass', sql.VarChar, String(password).trim());
    ins.input('role', sql.VarChar, String(role).trim());
    await ins.query(`
      INSERT INTO [dbo].[M5_Users] ([Nom],[Prenom],[Mlle],[Password],[Role])
      VALUES (@nom, @prenom, @mlle, @pass, @role);
    `);

    return res.status(201).json({
      matricule: mlleCanon,
      firstName: String(firstName).trim(),
      lastName: String(lastName).trim(),
      role: String(role).trim(),
      password: String(password).trim(),
      declaredCount: 0,
      correctedCount: 0
    });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: 'Failed to create user' });
  }
});

/* ---------------------------------------------------------------
DELETE /users
--------------------------------------------------------------- */
app.delete('/users', async (req, res) => {
  try {
    const pool = await getPool();

    const matricule = req.body?.matricule || req.query?.matricule;
    const code = req.body?.code || req.query?.code;

    if (!matricule || !code) {
      return res.status(400).json({ error: 'matricule and code are required.' });
    }

    const ADMIN_DELETE_CODE = '0000';
    if (String(code).trim() !== ADMIN_DELETE_CODE) {
      return res.status(401).json({ error: 'Invalid admin code.' });
    }

    const check = pool.request();
    check.input('mlle', sql.VarChar, String(matricule).trim());
    const user = await check.query(`
      SELECT * FROM [dbo].[M5_Users]
      WHERE LTRIM(RTRIM([Mlle])) = LTRIM(RTRIM(@mlle))
    `);

    if (!user.recordset.length) {
      return res.status(404).json({ error: 'User not found.' });
    }

    const del = pool.request();
    del.input('mlle', sql.VarChar, String(matricule).trim());
    await del.query(`
      DELETE FROM [dbo].[M5_Users]
      WHERE LTRIM(RTRIM([Mlle])) = LTRIM(RTRIM(@mlle))
    `);

    res.json({ success: true, message: 'User deleted successfully.' });
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: 'Failed to delete user.' });
  }
});

/* ---------------------------------------------------------------
PUT /users/:matricule
--------------------------------------------------------------- */
app.put('/users/:matricule', async (req, res) => {
  try {
    const pool = await getPool();
    const matricule = String(req.params.matricule || '').trim();

    if (!matricule) {
      return res.status(400).json({ error: 'Matricule is required in URL.' });
    }

    const { firstName, lastName, role, password } = req.body || {};

    if (!firstName && !lastName && !role && !password) {
      return res.status(400).json({ error: 'Nothing to update.' });
    }

    const check = pool.request();
    check.input('mlle', sql.VarChar, matricule);
    const existing = await check.query(`
      SELECT * FROM [dbo].[M5_Users]
      WHERE LTRIM(RTRIM([Mlle])) = LTRIM(RTRIM(@mlle))
    `);

    if (!existing.recordset.length) {
      return res.status(404).json({ error: 'User not found.' });
    }

    const fields = [];
    const updateReq = pool.request();

    if (firstName) {
      updateReq.input('prenom', sql.NVarChar, String(firstName).trim());
      fields.push('[Prenom] = @prenom');
    }
    if (lastName) {
      updateReq.input('nom', sql.NVarChar, String(lastName).trim());
      fields.push('[Nom] = @nom');
    }
    if (role) {
      updateReq.input('role', sql.VarChar, String(role).trim());
      fields.push('[Role] = @role');
    }
    if (password) {
      updateReq.input('pass', sql.VarChar, String(password).trim());
      fields.push('[Password] = @pass');
    }

    if (!fields.length) {
      return res.status(400).json({ error: 'No valid fields to update.' });
    }

    updateReq.input('mlle', sql.VarChar, matricule);

    const sqlQuery = `
      UPDATE [dbo].[M5_Users]
      SET ${fields.join(', ')}
      WHERE LTRIM(RTRIM([Mlle])) = LTRIM(RTRIM(@mlle));
    `;
    await updateReq.query(sqlQuery);

    const resReq = pool.request();
    resReq.input('mlle', sql.VarChar, matricule);
    const result = await resReq.query(`
      SELECT 
        [Nom],[Prenom],[Mlle],[Password],[Role],
        (SELECT COUNT(*) FROM [dbo].[M5_Urgent] WHERE LTRIM(RTRIM([Declarer_Par])) = LTRIM(RTRIM(u.[Mlle]))) AS declaredCount,
        (SELECT COUNT(*) FROM [dbo].[M5_Urgent] WHERE LTRIM(RTRIM([Corriger_Par])) = LTRIM(RTRIM(u.[Mlle]))) AS correctedCount
      FROM [dbo].[M5_Users] AS u
      WHERE LTRIM(RTRIM(u.[Mlle])) = LTRIM(RTRIM(@mlle));
    `);

    const r = result.recordset[0];
    res.json({
      matricule: r.Mlle?.trim(),
      firstName: r.Prenom?.trim(),
      lastName: r.Nom?.trim(),
      role: r.Role?.trim(),
      password: r.Password?.trim(),
      declaredCount: r.declaredCount ?? 0,
      correctedCount: r.correctedCount ?? 0
    });
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: 'Failed to update user.' });
  }
});

/* ---------------------------------------------------------------
GET /wires
--------------------------------------------------------------- */
app.get('/wires', async (req, res) => {
  try {
    const pool = await getPool();
    const conditions = [];
    const request = pool.request();

    if (ok(req.query.type)) {
      const t = String(req.query.type).toUpperCase().trim(); // COUPE / TWIST
      if (t === 'COUPE' || t === 'TWIST') {
        request.input('type', sql.VarChar, t);
        conditions.push('UPPER([Type]) = @type');
      }
    }

    if (ok(req.query.machine)) {
      request.input('mach', sql.VarChar, req.query.machine.trim());
      conditions.push('[Machine] = @mach');
    }

    if (ok(req.query.q)) {
      const like = `%${req.query.q.trim()}%`;
      request.input('q1', sql.VarChar, like);
      request.input('q2', sql.VarChar, like);
      conditions.push('([Unico] LIKE @q1 OR [Emplacement] LIKE @q2)');
    }

    const whereSql = conditions.length ? `WHERE ${conditions.join(' AND ')}` : '';

    const rows = await request.query(`
      SELECT [Unico], [Projet], [Emplacement], [Qte_Pq], [Machine], [Type]
      FROM [dbo].[M5_Wires]
      ${whereSql}
      ORDER BY [Unico] ASC;
    `);

    const results = rows.recordset.map(r => ({
      unico: r.Unico?.trim() || null,
      projet: r.Projet?.trim?.() || r.Projet || null,
      emplacement: r.Emplacement?.trim() || null,
      qte_pq: r.Qte_Pq ?? null,
      machine: r.Machine?.trim() || null,
      type: r.Type?.trim() || null
    }));

    res.json({ count: results.length, results });
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: 'Failed to fetch wires' });
  }
});

// health
app.get('/', (_, res) => res.json({ ok: true }));

/* ===================== Start server ===================== */
// app.listen(...)  ➜  server.listen(...)
server.listen(PORT, HOST, () => {
  console.log(`✅ API + WS running → http://${HOST}:${PORT}`);
});
